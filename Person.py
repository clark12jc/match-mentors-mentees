class Person:
    def __init__(self, first, last, location, capability, market, email, is_mentor):
        self.first = first.strip()
        self.last = last.strip()
        self.full_name = self.first + ' ' + self.last
        self.location = location.strip()
        self.capability = capability.strip()
        self.market = market.strip()
        self.email = email
        self.is_mentor = is_mentor
        self.has_match = False

    def __eq__(self, other):
        return (
            self.first.lower() == other.first.lower() and
            self.last.lower() == other.last.lower() and
            self.email.lower() == other.email.lower())

    def to_dict(self):
        return {
            'Full Name': self.full_name
        }
